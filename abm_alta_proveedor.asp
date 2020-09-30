<%@ LANGUAGE=VBScript %> 
<!--#INCLUDE file="library/PageOpen.asp"-->
<%

If (Session("Login") = 1) Then
else
	Response.Redirect "Mensaje.asp?Msg=No esta logueado en el sistema.&BackToURL=default.asp"
End If

ValidUserAction "ABM", "ABM"

set Rs_Pais = Server.CreateObject("ADODB.Recordset")
Rs_Pais.ActiveConnection = MM_HelpDesk_STRING
Rs_Pais.Source = "{call dbo.SPU_Pais_VE(" & 0 & "," & 1 & ")}"
Rs_Pais.Open()

If Request.Form ("idpais") <> 0 Then
  set RS_Ciudad = Server.CreateObject("ADODB.Recordset")
  RS_Ciudad.ActiveConnection = MM_HelpDesk_STRING
  RS_Ciudad.Source = "{call dbo.SPU_IDCiudadXPais_L(" + Replace(Request.Form("idpais"),"'","''") + ")}"
  RS_Ciudad.Open()
End if

%>

<script language="JavaScript">
function submitfrm()
{

	iErr = 0;
	if (frmaltaproveedor.proveedor.value == "")
	{
		alert ('El campo razon social es obligatorio.!!');
		iErr = 1;
	}
	if (frmaltaproveedor.idpais.value == "0")
	{
		alert ('El campo pais es obligatorio.!!');
		iErr = 1;
	}
	if (frmaltaproveedor.idCiudad.value == "0")
	{
		alert ('El campo ciudad es obligatorio.!!');
		iErr = 1;
	}

	if (iErr == 0)
	{
		frmaltaproveedor.action="vbsproveedor.asp";
		frmaltaproveedor.submit();
	}
}
</script>
          <table class="TopMenuArea" id="PageBorder" border="0" cellspacing="0" cellpadding="0" >
            <tr> 
              <td class="Titulo">Alta de Proveedores</td>
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
                <td>Ingrese los datos correspondientes al nuevo proveedor.
                </td>
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
                      <td width="17%">Razon Social(*):</td>
                      <td width="28%"> 
                        <input name="proveedor" type="text" class="FormTable" id="proveedor" size="20" value="<%If (Len(Request.Form("proveedor"))<>0) Then Response.Write Request.Form("proveedor")%>">
                      </td>
                      <td width="17%">Direccion:</td>
                      <td > 
                        <input type="text" name="direccion" size="30" class="FormTable" value="<%If Len(Request.Form("direccion"))<>0 Then Response.Write Request.Form("direccion")%>">
                      </td>
                    </tr>
                  </table>
                   
              <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr> 
                      <td width="17%">Codigo Postal:</td>
                      <td width="28%"> 
                        <input name="codigopostal" type="text" class="FormTable" id="codigopostal" size="10" value="<%If Len(Request.Form("codigopostal"))<>0 Then Response.Write Request.Form("codigopostal")%>">
                      </td>
                      <td width="17%">Telefono:</td>
                      <td > 
                        <input type="text" name="telefono" size="20" class="FormTable" value="<%If Len(Request.Form("telefono"))<>0 Then Response.Write Request.Form("telefono")%>">
                      </td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr> 
                      <td width="17%">Pagina web:</td>
                      <td width="28%"> 
                        <input type="text" name="paginaweb" size="20" class="FormTable" value="<%If Len(Request.Form("paginaweb"))<>0 then Response.Write Request.Form("paginaweb")%>">
                      </td>
                      <td width="17%">E-Mail:</td>
                      <td width="42%"> 
                        <input name="email" type="text" class="FormTable" id="email" size="20" onBlur="MM_validateForm('email','','NisEmail');return document.MM_returnValue" value="<%If Len(Request.Form("email"))<>0 Then Response.Write Request.Form("email")%>">
                      </td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr> 
                      <td width="17%">Pais(*):</td>
                      <td width="28%"><select name="idpais" size="1" class="FormTable" id="idpais" onChange="JavaScript:reloadpage(frmaltaproveedor);">
                          <option value="0">----</option>
                        <%
						While (NOT Rs_Pais.EOF)
						%>
                          <option value="<%=(Rs_Pais.Fields.Item("IDPais").Value)%>"
                               <%if (CStr(Request.Form("idpais"))=CStr(RS_Pais("IDPais"))) Then Response.Write("SELECTED")%>>
                          <%=(Rs_Pais("Pais"))%></option>
                        <%
						  Rs_Pais.MoveNext()
						Wend
						If (Rs_Pais.CursorType > 0) Then
						  Rs_Pais.MoveFirst
						Else
						  Rs_Pais.Requery
						End If
						%>
                        </select></td>
                      <td width="17%">Ciudad(*):</td>
                      <td width="42%"> 
                        <select name="idCiudad" size="1" class="FormTable" id="idCiudad">
                          <option value="0">----</option>
						<%
                        If  Request.Form ("idpais") <> 0 Then
							While (NOT Rs_Ciudad.EOF)
							%>
							  <option value="<%=(Rs_Ciudad.Fields.Item("IDCiudad").Value)%>"><%=(Rs_Ciudad.Fields.Item("Ciudad").Value)%></option>
							<%
							  Rs_Ciudad.MoveNext()
							Wend
							If (Rs_Ciudad.CursorType > 0) Then
							  Rs_Ciudad.MoveFirst
							Else
							  Rs_Ciudad.Requery
							End If
						End If
						%>
                        </select></td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="FormTable">
                    <tr> 
                      <td width="17%">Contacto:</td>
                      <td width="28%"> 
                        <input type="text" name="contacto" size="20" class="FormTable" value="<%If Len(Request.Form("contacto"))<>0 Then Response.Write Request.Form("contacto")%>">
                      </td>
                      <td width="17%" > 
                        Interno:</td>
                      <td > 
                        <input type="text" name="interno" size="10" class="FormTable" value="<%If Len(Request.Form("interno"))<>0 Then Response.Write Request.Form("interno")%>">
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
                          <input type="button" value="Ingresar" name="Ingresar" class="FormButton" onClick="JavaScript:submitfrm()">
                          <input type="reset" value="Limpiar" name="Limpiar" class="FormButton">
                      </td>
                    </tr>
                  </table>
                  
                  </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </form>
       
      <!--#INCLUDE file="library/PageClose.asp"-->
<%
Rs_Pais.Close()
If Request.Form ("idpais") <> 0 Then
   Rs_Ciudad.Close()
End If
%>